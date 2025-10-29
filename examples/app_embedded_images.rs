// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of embedding images into a worksheet cells using `rust_xlsxwriter`.
//! This image scales to size of the cell and moves with it.
//!
//! This approach can be useful if you are building up a spreadsheet of products
//! with a column of images for each product.
//!
//! This is the equivalent of Excel's menu option to insert an image using the
//! option to "Place in Cell" which is only available in Excel 365 versions from
//! 2023 onwards. For older versions of Excel a `#VALUE!` error is displayed.
//!
use rust_xlsxwriter::{Format, FormatAlign, Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create some formats to use in the example.
    let vertical_center = Format::new().set_align(FormatAlign::VerticalCenter);
    let center = Format::new().set_align(FormatAlign::Center);

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a new image object.
    let image = Image::new("examples/rust_logo.png")?;

    // Widen the first column to make the captions clearer.
    worksheet.set_column_width(0, 30)?;

    // Change cell widths/heights to demonstrate the image differences.
    worksheet.set_column_width(1, 14)?;
    worksheet.set_row_height(1, 60)?;
    worksheet.set_row_height(3, 60)?;
    worksheet.set_row_height(5, 90)?;

    // Embed an image in a cell. The height and width scale automatically.
    worksheet.write_with_format(1, 0, "Embed image in cell:", &vertical_center)?;
    worksheet.embed_image(1, 1, &image)?;

    // Embed and center an image in a cell.
    worksheet.write_with_format(3, 0, "Embed and center image:", &vertical_center)?;
    worksheet.embed_image_with_format(3, 1, &image, &center)?;

    // Embed an image in a larger cell.
    worksheet.write_with_format(5, 0, "Embed image in larger cell:", &vertical_center)?;
    worksheet.embed_image(5, 1, &image)?;

    // Save the file to disk.
    workbook.save("embedded_images.xlsx")?;

    Ok(())
}
