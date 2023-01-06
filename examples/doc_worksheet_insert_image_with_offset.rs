// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! This example shows how to add an image to a worksheet at an offset within
//! the cell.

use rust_xlsxwriter::{Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a new image object.
    let image = Image::new("examples/rust_logo.png")?;

    // Insert the image at an offset.
    worksheet.insert_image_with_offset(1, 2, &image, 10, 5)?;

    // Save the file to disk.
    workbook.save("image.xlsx")?;

    Ok(())
}
