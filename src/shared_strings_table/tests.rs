// Shared Strings Table unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod shared_strings_table_tests {

    use crate::shared_strings_table::SharedStringsTable;

    #[test]
    fn test_shared_string_table() {
        let mut string_table = SharedStringsTable::new();

        let index = string_table.shared_string_index("neptune".into());
        assert_eq!(index, 0);

        let index = string_table.shared_string_index("neptune".into());
        assert_eq!(index, 0);

        let index = string_table.shared_string_index("neptune".into());
        assert_eq!(index, 0);

        let index = string_table.shared_string_index("mars".into());
        assert_eq!(index, 1);

        let index = string_table.shared_string_index("venus".into());
        assert_eq!(index, 2);

        let index = string_table.shared_string_index("mars".into());

        assert_eq!(index, 1);

        let index = string_table.shared_string_index("venus".into());
        assert_eq!(index, 2);
    }
}
