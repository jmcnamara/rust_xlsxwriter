// Shared Strings unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod shared_strings_tests {

    use std::sync::{Arc, Mutex};

    use crate::shared_strings::SharedStrings;
    use crate::shared_strings_table::SharedStringsTable;
    use crate::test_functions::xml_to_vec;
    use crate::xmlwriter;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_shared_string_table() {
        let mut string_table = SharedStringsTable::new();

        let mut shared_strings = SharedStrings::new();

        string_table.shared_string_index("neptune".into());
        string_table.shared_string_index("neptune".into());
        string_table.shared_string_index("neptune".into());
        string_table.shared_string_index("neptune".into());
        string_table.shared_string_index("mars".into());
        string_table.shared_string_index("venus".into());
        string_table.shared_string_index("mars".into());

        let string_table = Arc::new(Mutex::new(string_table));

        shared_strings.assemble_xml_file(string_table);

        let got = xmlwriter::cursor_to_str(&shared_strings.writer);
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

        string_table.shared_string_index("abcdefg".into());
        string_table.shared_string_index("   abcdefg".into());
        string_table.shared_string_index("abcdefg   ".into());

        let string_table = Arc::new(Mutex::new(string_table));

        shared_strings.assemble_xml_file(string_table);

        let got = xmlwriter::cursor_to_str(&shared_strings.writer);
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
