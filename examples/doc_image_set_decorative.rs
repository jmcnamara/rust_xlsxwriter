// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! This example shows how to create an image object and set the decorative
//! property to indicate the it doesn't contain useful visual information. This
//! is used to improve the accessibility of visual elements.

use rust_xlsxwriter::{Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a new image object.
    let image = Image::new("examples/rust_logo.png")?.set_decorative(true);

    // Insert the image.
    worksheet.insert_image(1, 2, &image)?;

    // Save the file to disk.
    workbook.save("image.xlsx")?;

    Ok(())
}
