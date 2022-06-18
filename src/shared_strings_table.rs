// shared_strings_table - A module for storing Excel shared strings.
//
// Excel doesn't store strings directly in the worksheet?.xml files. Instead
// it stores them in a hash table with an index number, based on the order of
// writing and writes the index to the worksheet instead.
//
// SPDX-License-Identifier: MIT
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use std::collections::HashMap;

//
// A metadata stuct to store Excel unique strings between worksheets.
//
pub struct SharedStringsTable {
    pub count: u32,
    pub unique_count: u32,
    pub strings: HashMap<String, u32>,
}

impl SharedStringsTable {
    // Create a new struct to to track Excel shared strings between worksheets.
    pub fn new() -> SharedStringsTable {
        SharedStringsTable {
            count: 0,
            unique_count: 0,
            strings: HashMap::new(),
        }
    }

    // Get the index of the string in the Shared String table.
    pub fn get_shared_string_index(&mut self, key: &str) -> u32 {
        match self.strings.get(key) {
            Some(value) => {
                self.count += 1;
                *value
            }
            None => {
                let index = self.unique_count;
                self.strings.insert(key.to_string(), self.unique_count);
                self.count += 1;
                self.unique_count += 1;
                index
            }
        }
    }
}

#[cfg(test)]
mod tests {

    use super::SharedStringsTable;

    #[test]
    fn test_shared_string_table() {
        let mut string_table = SharedStringsTable::new();

        let index = string_table.get_shared_string_index("neptune");
        assert_eq!(index, 0);

        let index = string_table.get_shared_string_index("neptune");
        assert_eq!(index, 0);

        let index = string_table.get_shared_string_index("neptune");
        assert_eq!(index, 0);

        let index = string_table.get_shared_string_index("mars");
        assert_eq!(index, 1);

        let index = string_table.get_shared_string_index("venus");
        assert_eq!(index, 2);

        let index = string_table.get_shared_string_index("mars");

        assert_eq!(index, 1);

        let index = string_table.get_shared_string_index("venus");
        assert_eq!(index, 2);
    }
}
