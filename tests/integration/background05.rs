// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{Image, Workbook, XlsxError};

// Create rust_xlsxwriter file to compare against Excel file.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();
    let image = Image::new("tests/input/images/logo.jpg")?;
    worksheet.insert_background_image(&image);

    let worksheet = workbook.add_worksheet();
    let image = Image::new("tests/input/images/red.jpg")?;
    worksheet.insert_background_image(&image);

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_background05() {
    let test_runner = common::TestRunner::new()
        .set_name("background05")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
