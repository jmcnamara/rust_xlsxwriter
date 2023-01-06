// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! This example shows how to create an image object and use it to insert the
//! image into a worksheet.

use rust_xlsxwriter::{Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a new image object.
    let image = Image::new("examples/rust_logo.png")?;

    // Insert the image.
    worksheet.insert_image(1, 2, &image)?;

    // Save the file to disk.
    workbook.save("image.xlsx")?;

    Ok(())
}
