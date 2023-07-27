// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{Workbook, XlsxError};

// Test the creation of a simple rust_xlsxwriter file with repeat rows/cols.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet1 = workbook.add_worksheet();
    worksheet1.write_string(0, 0, "Foo")?;
    worksheet1.set_repeat_rows(0, 0)?;

    let worksheet2 = workbook.add_worksheet();
    worksheet2.set_portrait();

    let worksheet3 = workbook.add_worksheet();
    worksheet3.set_repeat_rows(2, 3)?;
    worksheet3.set_repeat_columns(1, 5)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_repeat05() {
    let test_runner = common::TestRunner::new()
        .set_name("repeat05")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}