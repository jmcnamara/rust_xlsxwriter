// Test case that compares a file generated by rust_xlsxwriter with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use rust_xlsxwriter::{Image, Workbook, XlsxError};

mod common;

// Test to demonstrate handling duplicate images in worksheets.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let mut image = Image::new("tests/input/images/red.png")?;
    image.set_alt_text("red.png");

    let worksheet1 = workbook.add_worksheet();
    worksheet1.insert_image(8, 4, &image)?;

    let worksheet2 = workbook.add_worksheet();
    worksheet2.insert_image(8, 4, &image)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_image48() {
    let test_runner = common::TestRunner::new()
        .set_name("image48")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
