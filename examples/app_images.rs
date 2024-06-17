// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of inserting images into a worksheet using rust_xlsxwriter.

use rust_xlsxwriter::{Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Widen the first column to make the text clearer.
    worksheet.set_column_width(0, 30)?;

    // Create a new image object.
    let mut image = Image::new("examples/rust_logo.png")?;

    // Insert the image.
    worksheet.write_string(0, 0, "Insert an image in a cell:")?;
    worksheet.insert_image(0, 1, &image)?;

    // Insert an image offset in the cell.
    worksheet.write_string(7, 0, "Insert an image with an offset:")?;
    worksheet.insert_image_with_offset(7, 1, &image, 5, 5)?;

    // Insert an image with scaling.
    worksheet.write_string(15, 0, "Insert a scaled image:")?;
    image = image.set_scale_width(0.75).set_scale_height(0.75);
    worksheet.insert_image(15, 1, &image)?;

    // Save the file to disk.
    workbook.save("images.xlsx")?;

    Ok(())
}
