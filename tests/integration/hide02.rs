// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};

// Test to demonstrate very hidden worksheets.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let mut worksheet1 = Worksheet::new();
    let mut worksheet2 = Worksheet::new();
    let worksheet3 = Worksheet::new();

    // This should be ignored since there is no other active worksheet.
    worksheet1.set_very_hidden(true);

    // This should be set.
    worksheet2.set_very_hidden(true);

    workbook.push_worksheet(worksheet1);
    workbook.push_worksheet(worksheet2);
    workbook.push_worksheet(worksheet3);

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_hide02() {
    let test_runner = common::TestRunner::new()
        .set_name("hide02")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
