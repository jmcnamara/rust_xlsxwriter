// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{Image, Workbook, XlsxError};

// Test to demonstrate adding images to worksheets.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    let image = Image::new("tests/input/images/logo.png")?.set_alt_text("logo.png");

    worksheet.insert_image(1, 2, &image)?; // Test double adding.
    worksheet.insert_image_with_offset(1, 2, &image, 8, 5)?;

    workbook.save(filename)?;
    workbook.save(filename)?; // Test double saving.

    Ok(())
}

#[test]
fn test_image11() {
    let test_runner = common::TestRunner::new()
        .set_name("image11")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
