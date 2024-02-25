// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of embedding images into a worksheet cells using rust_xlsxwriter.
//! This image scales to size of the cell and moves with it.
//!
//! This approach can be useful if you are building up a spreadsheet of products
//! with a column of images for each product.
//!
//! This is the equivalent of Excel's menu option to insert an image using the
//! option to "Place in Cell" which is only available in Excel 365 versions from
//! 2023 onwards. For older versions of Excel a `#VALUE!` error is displayed.
//!
use rust_xlsxwriter::{Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a new image object.
    let image = Image::new("examples/rust_logo.png")?;

    // Widen the first column to make the caption clearer.
    worksheet.set_column_width(0, 30)?;
    worksheet.write(0, 0, "Embed images that scale to the cell size")?;

    // Change cell widths/heights to demonstrate the image differences.
    worksheet.set_column_width(1, 14)?;
    worksheet.set_row_height(1, 60)?;
    worksheet.set_row_height(3, 90)?;

    // Embed the images in cells of different widths/heights.
    worksheet.embed_image(1, 1, &image)?;
    worksheet.embed_image(3, 1, &image)?;

    // Save the file to disk.
    workbook.save("embedded_images.xlsx")?;

    Ok(())
}
