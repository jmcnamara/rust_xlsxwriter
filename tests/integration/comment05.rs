// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{Note, Workbook, XlsxError};

// Create a rust_xlsxwriter file to compare against an Excel file.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet1 = workbook.add_worksheet();
    worksheet1.set_default_note_author("John");

    let note = Note::new("Some text").add_author_prefix(false);

    for row in 0..=127 {
        for col in 0..=15 {
            worksheet1.insert_note(row, col, &note)?;
        }
    }

    let _worksheet2 = workbook.add_worksheet();

    let worksheet3 = workbook.add_worksheet();
    worksheet3.set_default_note_author("John");

    let note = Note::new("More text").add_author_prefix(false);
    worksheet3.insert_note(0, 0, &note)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_comment05() {
    let test_runner = common::TestRunner::new()
        .set_name("comment05")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
