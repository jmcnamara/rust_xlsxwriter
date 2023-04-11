// shared_strings_table - A module for storing Excel shared strings.
//
// Excel doesn't store strings directly in the worksheet?.xml files. Instead
// it stores them in a hash table with an index number, based on the order of
// writing and writes the index to the worksheet instead.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use std::{collections::HashMap, rc::Rc};

//
// A metadata struct to store Excel unique strings between worksheets.
//
pub struct SharedStringsTable {
    pub count: u32,
    pub unique_count: u32,
    pub strings: HashMap<Rc<str>, u32>,
}

impl SharedStringsTable {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new struct to to track Excel shared strings between worksheets.
    pub(crate) fn new() -> SharedStringsTable {
        SharedStringsTable {
            count: 0,
            unique_count: 0,
            strings: HashMap::new(),
        }
    }

    // Get the index of the string in the Shared String table.
    pub(crate) fn shared_string_index(&mut self, key: Rc<str>) -> u32 {
        let index = *self.strings.entry(key).or_insert_with(|| {
            let index = self.unique_count;
            self.unique_count += 1;
            index
        });
        self.count += 1;
        index
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

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
