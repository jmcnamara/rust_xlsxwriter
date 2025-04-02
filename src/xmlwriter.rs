// xmlwriter - a module for writing XML in the same format and with the same
// escaping as used by Excel in xlsx xml files. This is a common set of
// functionality for all of the other xml writing structs.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

mod tests;

use std::borrow::Cow;
use std::io::{Cursor, Write};
use std::str;

pub(crate) const XML_WRITE_ERROR: &str = "Couldn't write to xml file";
const UNICODE_ESCAPE_LENGTH: usize = 7; // Length of _xHHHH_.

// -----------------------------------------------------------------------
// XML Writing functions.
// -----------------------------------------------------------------------

// Write the XML declaration for a file.
pub(crate) fn xml_declaration<W: Write>(mut writer: W) {
    writer
        .write_all(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
        .expect(XML_WRITE_ERROR);
}

// Write an XML start tag that has no attributes.
pub(crate) fn xml_start_tag_only<W: Write>(writer: &mut W, tag: &str) {
    write!(writer, "<{tag}>").expect(XML_WRITE_ERROR);
}

// Write an XML start tag with specified attributes.
pub(crate) fn xml_start_tag<W, T>(writer: &mut W, tag: &str, attributes: &[T])
where
    W: Write,
    T: IntoAttribute,
{
    write!(writer, "<{tag}").expect(XML_WRITE_ERROR);

    for attribute in attributes {
        attribute.write_to(writer);
    }

    writer.write_all(b">").expect(XML_WRITE_ERROR);
}

// Write the closing tag for an XML element.
pub(crate) fn xml_end_tag<W: Write>(writer: &mut W, tag: &str) {
    write!(writer, "</{tag}>").expect(XML_WRITE_ERROR);
}

// Write an empty XML tag that has no attributes.
pub(crate) fn xml_empty_tag_only<W: Write>(writer: &mut W, tag: &str) {
    write!(writer, "<{tag}/>").expect(XML_WRITE_ERROR);
}

// Write an empty XML tag with specified attributes.
pub(crate) fn xml_empty_tag<W, T>(writer: &mut W, tag: &str, attributes: &[T])
where
    W: Write,
    T: IntoAttribute,
{
    write!(writer, "<{tag}").expect(XML_WRITE_ERROR);

    for attribute in attributes {
        attribute.write_to(writer);
    }

    writer.write_all(b"/>").expect(XML_WRITE_ERROR);
}

// Write an XML element with data but no attributes.
pub(crate) fn xml_data_element_only<W: Write>(writer: &mut W, tag: &str, data: &str) {
    write!(writer, "<{}>{}</{}>", tag, escape_xml_data(data), tag).expect(XML_WRITE_ERROR);
}

// Write an XML element with data and specified attributes.
pub(crate) fn xml_data_element<W, T>(writer: &mut W, tag: &str, data: &str, attributes: &[T])
where
    W: Write,
    T: IntoAttribute,
{
    write!(writer, "<{tag}").expect(XML_WRITE_ERROR);

    for attribute in attributes {
        attribute.write_to(writer);
    }

    write!(writer, ">{}</{}>", escape_xml_data(data), tag).expect(XML_WRITE_ERROR);
}

// Optimized writer for <si> elements in shared strings.
pub(crate) fn xml_si_element<W: Write>(writer: &mut W, string: &str, preserve_whitespace: bool) {
    if preserve_whitespace {
        write!(
            writer,
            r#"<si><t xml:space="preserve">{}</t></si>"#,
            escape_xml_data(&escape_xml_escapes(string))
        )
        .expect(XML_WRITE_ERROR);
    } else {
        write!(
            writer,
            "<si><t>{}</t></si>",
            escape_xml_data(&escape_xml_escapes(string))
        )
        .expect(XML_WRITE_ERROR);
    }
}

// Write an <si> element for rich strings.
pub(crate) fn xml_rich_si_element<W: Write>(writer: &mut W, string: &str) {
    write!(writer, "<si>{string}</si>").expect(XML_WRITE_ERROR);
}

// Write the theme string into the theme file.
pub(crate) fn xml_theme<W: Write>(writer: &mut W, theme: &str) {
    writeln!(writer, "{theme}").expect(XML_WRITE_ERROR);
}

// Write a string after escaping XML characters.
pub(crate) fn xml_raw_string<W: Write>(writer: &mut W, data: &str) {
    writer.write_all(data.as_bytes()).expect(XML_WRITE_ERROR);
}

// Escape special characters in XML attributes.
pub(crate) fn escape_attributes(attribute: &str) -> Cow<str> {
    escape_string(attribute, match_attribute_html_char)
}

// Escape special characters in the data sections of XML tags.
pub(crate) fn escape_xml_data(data: &str) -> Cow<str> {
    escape_string(data, match_xml_char)
}

// Escape non-URL-safe characters in a hyperlink or URL.
pub(crate) fn escape_url(data: &str) -> Cow<str> {
    escape_string(data, match_url_char)
}

// -----------------------------------------------------------------------
// Helper functions for XML, primarily for string escaping.
// -----------------------------------------------------------------------

// Helper function to read XML data from a cursor for testing.
#[allow(dead_code)]
pub(crate) fn cursor_to_str(cursor: &Cursor<Vec<u8>>) -> &str {
    str::from_utf8(cursor.get_ref()).unwrap()
}

pub(crate) fn cursor_to_string(cursor: &Cursor<Vec<u8>>) -> String {
    str::from_utf8(cursor.get_ref()).unwrap().to_string()
}

// Reset the memory cursor buffer, typically used between saves.
pub(crate) fn reset(cursor: &mut Cursor<Vec<u8>>) {
    cursor.get_mut().clear();
    cursor.set_position(0);
}

// Matching function used by escape_attributes().
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

// Matching function used by escape_xml_data().
//
// Note: This differs from match_attribute_html_char() because Excel does not
// escape double quotes or newlines.
//
// To mimic Excel, escape control and non-printing characters in the range
// '\x00' - '\x1F'.
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

// Match the URL characters that Excel escapes.
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

// Generic escape function that uses a function pointer for the required handler.
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

// Excel escapes control characters with _xHHHH_ (see match_xml_char() above).
// As a result, it also escapes any literal strings of that type by encoding the
// leading underscore. For example, "_x0000_" becomes "_x005F_x0000_".
pub(crate) fn escape_xml_escapes(original: &str) -> Cow<str> {
    if !original.contains("_x") {
        return Cow::Borrowed(original);
    }

    let string_end = original.len();
    let mut escaped_string = original.to_string();

    // Match from right so we can escape target string at the same indices.
    let matches: Vec<_> = original.rmatch_indices("_x").collect();

    for (index, _) in matches {
        if index + UNICODE_ESCAPE_LENGTH > string_end {
            continue;
        }

        // Ensure that the end index is at a character boundary.
        if !original.is_char_boundary(index + 6) {
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

// -----------------------------------------------------------------------
// Trait to write attribute tuple values to an XML file.
// -----------------------------------------------------------------------

pub(crate) trait IntoAttribute {
    fn write_to<W>(&self, writer: &mut W)
    where
        W: Write;
}

impl IntoAttribute for (&str, &str) {
    fn write_to<W>(&self, writer: &mut W)
    where
        W: Write,
    {
        write!(writer, r#" {}="{}""#, self.0, escape_attributes(self.1)).expect(XML_WRITE_ERROR);
    }
}

impl IntoAttribute for (&str, String) {
    fn write_to<W>(&self, writer: &mut W)
    where
        W: Write,
    {
        write!(writer, r#" {}="{}""#, self.0, escape_attributes(&self.1)).expect(XML_WRITE_ERROR);
    }
}
