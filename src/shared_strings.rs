// shared_strings - A module for creating the Excel sharedStrings.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use itertools::Itertools;

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

    //  Assemble and write the XML file.
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
        let attributes = vec![("xmlns", xmls), ("count", count), ("uniqueCount", unique)];

        self.writer.xml_start_tag_attr("sst", &attributes);
    }

    // Write the sst string elements.
    fn write_sst_strings(&mut self, string_table: &SharedStringsTable) {
        for (string, _) in string_table.strings.iter().sorted_by_key(|x| x.1) {
            let preserve_whitespace =
                string.starts_with(['\t', '\n', ' ']) || string.ends_with(['\t', '\n', ' ']);

            if string.starts_with("<r>") && string.ends_with("</r>") {
                self.writer.xml_rich_si_element(string);
            } else {
                self.writer.xml_si_element(string, preserve_whitespace);
            }
        }
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::shared_strings::SharedStrings;
    use crate::shared_strings_table::SharedStringsTable;
    use crate::test_functions::xml_to_vec;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_shared_string_table() {
        let mut string_table = SharedStringsTable::new();

        let mut shared_strings = SharedStrings::new();

        string_table.shared_string_index("neptune");
        string_table.shared_string_index("neptune");
        string_table.shared_string_index("neptune");
        string_table.shared_string_index("neptune");
        string_table.shared_string_index("mars");
        string_table.shared_string_index("venus");
        string_table.shared_string_index("mars");

        shared_strings.assemble_xml_file(&string_table);

        let got = shared_strings.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="3">
                  <si>
                    <t>neptune</t>
                  </si>
                  <si>
                    <t>mars</t>
                  </si>
                  <si>
                    <t>venus</t>
                  </si>
                </sst>
                "#,
        );

        assert_eq!(expected, got);
    }

    #[test]
    fn test_shared_string_table_with_preserve() {
        let mut string_table = SharedStringsTable::new();

        let mut shared_strings = SharedStrings::new();

        string_table.shared_string_index("abcdefg");
        string_table.shared_string_index("   abcdefg");
        string_table.shared_string_index("abcdefg   ");

        shared_strings.assemble_xml_file(&string_table);

        let got = shared_strings.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
                <si>
                    <t>abcdefg</t>
                </si>
                <si>
                    <t xml:space="preserve">   abcdefg</t>
                </si>
                <si>
                    <t xml:space="preserve">abcdefg   </t>
                </si>
            </sst>
                "#,
        );

        assert_eq!(expected, got);
    }
}
