// shared_strings - A module for creating the Excel sharedStrings.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

mod tests;

use crate::shared_strings_table::SharedStringsTable;
use crate::xmlwriter::XMLWriter;

pub struct SharedStrings {
    pub(crate) writer: XMLWriter,
}

impl SharedStrings {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new SharedStrings struct.
    pub(crate) fn new() -> SharedStrings {
        let writer = XMLWriter::new();

        SharedStrings { writer }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self, string_table: &SharedStringsTable) {
        self.writer.xml_declaration();

        // Write the sst element.
        self.write_sst(string_table);

        // Write the sst strings.
        self.write_sst_strings(string_table);

        // Close the sst tag.
        self.writer.xml_end_tag("sst");
    }

    // Write the <sst> element.
    fn write_sst(&mut self, string_table: &SharedStringsTable) {
        let xmls = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string();
        let count = string_table.count.to_string();
        let unique = string_table.unique_count.to_string();
        let attributes = [("xmlns", xmls), ("count", count), ("uniqueCount", unique)];

        self.writer.xml_start_tag("sst", &attributes);
    }

    // Write the sst string elements.
    #[allow(clippy::from_iter_instead_of_collect)] // from_iter() is faster than collect() here.
    fn write_sst_strings(&mut self, string_table: &SharedStringsTable) {
        let mut insertion_order_strings = Vec::from_iter(string_table.strings.iter());
        insertion_order_strings.sort_by_key(|x| x.1);
        let whitespace = ['\t', '\n', ' '];

        for (string, _) in insertion_order_strings {
            let preserve_whitespace =
                string.starts_with(whitespace) || string.ends_with(whitespace);

            if string.starts_with("<r>") && string.ends_with("</r>") {
                self.writer.xml_rich_si_element(string);
            } else {
                self.writer.xml_si_element(string, preserve_whitespace);
            }
        }
    }
}
