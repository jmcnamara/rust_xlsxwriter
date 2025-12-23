// Test case for the image centering bug fix.
//
// Bug: Floating-point precision caused scaled_height to be slightly less than
// cell_height (e.g., 14.9999... as u32 = 14 != 15), which caused the wrong
// centering logic to be applied.
//
// Fix: Use .round() before casting to u32.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

use crate::common;
use rust_xlsxwriter::{Image, Workbook, XlsxError};

// Test the fix for the centering bug:
// Image 20x11 in cell 200x15 should be horizontally centered.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.set_column_width_pixels(0, 200)?;
    worksheet.set_row_height_pixels(0, 15)?;

    let image = Image::new("tests/input/images/red_20x11.png")?;

    worksheet.insert_image_fit_to_cell_centered(0, 0, &image)?;

    workbook.save(filename)?;

    Ok(())
}

#[test]
fn test_image_fit_to_cell_centering() {
    let test_runner = common::TestRunner::new()
        .set_name("image_fit_to_cell_centering")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
