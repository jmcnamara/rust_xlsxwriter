// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! This example shows how to create an image object and use it to insert the
//! image into a worksheet. The image in this case is scaled by setting the
//! height and width.

use rust_xlsxwriter::{Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a new image object and set the image logical sizes.
    let image = Image::new("examples/rust_logo.png")?
        .set_height(80)
        .set_width(80);

    // Insert the image.
    worksheet.insert_image(1, 2, &image)?;

    // Save the file to disk.
    workbook.save("image.xlsx")?;

    Ok(())
}
