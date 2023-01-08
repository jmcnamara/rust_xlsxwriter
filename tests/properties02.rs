// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use rust_xlsxwriter::{Properties, Workbook, XlsxError};

mod common;

// Test to demonstrate document properties.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let properties = Properties::new().set_hyperlink_base("C:\\");

    workbook.set_properties(&properties);

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_properties02() {
    let test_runner = common::TestRunner::new()
        .set_name("properties02")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
