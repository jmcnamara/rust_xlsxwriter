// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of inserting a background image into a worksheet using
//! `rust_xlsxwriter`.
//!
//! See also the `app_watermark.rs` example which shows how to set a watermark
//! via the header image of a worksheet. That is the way that the Microsoft
//! documentation recommends to set a watermark in Excel.

use rust_xlsxwriter::{Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // The image may not be visible unless the view is large.
    worksheet.write(0, 0, "Scroll down and right to see the background image")?;

    // Create a new image object.
    let image = Image::new("examples/watermark.png")?;

    // Insert the background image.
    worksheet.insert_background_image(&image);

    // Save the file to disk.
    workbook.save("background_image.xlsx")?;

    Ok(())
}
