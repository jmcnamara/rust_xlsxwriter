// shared_strings - A module for creating the Excel sharedStrings.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

mod tests;

use std::io::Cursor;
use std::sync::{Arc, Mutex};

use crate::shared_strings_table::SharedStringsTable;

use crate::xmlwriter::{
    xml_declaration, xml_end_tag, xml_rich_si_element, xml_si_element, xml_start_tag,
};

pub struct SharedStrings {
    pub(crate) writer: Cursor<Vec<u8>>,
}

impl SharedStrings {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new SharedStrings struct.
    pub(crate) fn new() -> SharedStrings {
        let writer = Cursor::new(Vec::with_capacity(2048));

        SharedStrings { writer }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and generate the XML file.
    pub(crate) fn assemble_xml_file(&mut self, string_table: Arc<Mutex<SharedStringsTable>>) {
        xml_declaration(&mut self.writer);

        // Write the <sst> element.
        self.write_sst(&string_table);

        // Write the <sst> strings.
        self.write_sst_strings(&string_table);

        // Close the <sst> tag.
        xml_end_tag(&mut self.writer, "sst");
    }

    // Write the <sst> element.
    fn write_sst(&mut self, string_table: &Arc<Mutex<SharedStringsTable>>) {
        let string_table = string_table.lock().unwrap();

        let xmls = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string();
        let count = string_table.count.to_string();
        let unique = string_table.unique_count.to_string();
        let attributes = [("xmlns", xmls), ("count", count), ("uniqueCount", unique)];

        xml_start_tag(&mut self.writer, "sst", &attributes);
    }

    // Write the <sst> string elements.
    #[allow(clippy::from_iter_instead_of_collect)] // from_iter() is faster than collect() here.
    fn write_sst_strings(&mut self, string_table: &Arc<Mutex<SharedStringsTable>>) {
        let string_table = string_table.lock().unwrap();

        let mut insertion_order_strings = Vec::from_iter(string_table.strings.iter());
        insertion_order_strings.sort_by_key(|x| x.1);
        let whitespace = ['\t', '\n', ' '];

        for (string, _) in insertion_order_strings {
            let preserve_whitespace =
                string.starts_with(whitespace) || string.ends_with(whitespace);

            // Check if the string is a rich text element.
            if string.starts_with("<r>") && string.ends_with("</r>") {
                xml_rich_si_element(&mut self.writer, string);
            } else {
                xml_si_element(&mut self.writer, string, preserve_whitespace);
            }
        }
    }
}
