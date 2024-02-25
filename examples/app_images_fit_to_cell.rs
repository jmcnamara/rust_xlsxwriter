// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of inserting images into a worksheet using rust_xlsxwriter so
//! that they are scaled to a cell. This approach can be useful if you are
//! building up a spreadsheet of products with a column of images for each
//! product.
//!
//! See also the `app_embedded_image.rs` example that shows a better approach
//! for newer versions of Excel.

use rust_xlsxwriter::{Format, FormatAlign, Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    let center = Format::new().set_align(FormatAlign::VerticalCenter);

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Widen the first column to make the text clearer.
    worksheet.set_column_width(0, 30)?;

    // Set larger cells to accommodate the images.
    worksheet.set_column_width_pixels(1, 200)?;
    worksheet.set_row_height_pixels(0, 140)?;
    worksheet.set_row_height_pixels(2, 140)?;
    worksheet.set_row_height_pixels(4, 140)?;

    // Create a new image object.
    let image = Image::new("examples/rust_logo.png")?;

    // Insert the image as standard, without scaling.
    worksheet.write_with_format(0, 0, "Unscaled image inserted into cell:", &center)?;
    worksheet.insert_image(0, 1, &image)?;

    // Insert the image and scale it to fit the entire cell.
    worksheet.write_with_format(2, 0, "Image scaled to fit cell:", &center)?;
    worksheet.insert_image_fit_to_cell(2, 1, &image, false)?;

    // Insert the image and scale it to the cell while maintaining the aspect ratio.
    // In this case it is scaled to the smaller of the width or height scales.
    worksheet.write_with_format(4, 0, "Image scaled with a fixed aspect ratio:", &center)?;
    worksheet.insert_image_fit_to_cell(4, 1, &image, true)?;

    // Save the file to disk.
    workbook.save("images_fit_to_cell.xlsx")?;

    Ok(())
}
