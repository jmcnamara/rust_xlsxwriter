// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! An example of scaling images to a fixed width and height. See also the
//! `worksheet.insert_image_fit_to_cell()` method.

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
    let mut image = Image::new("examples/rust_logo.png")?;

    // Insert the image as standard, without scaling.
    worksheet.write_with_format(0, 0, "Unscaled image inserted into cell:", &center)?;
    worksheet.insert_image(0, 1, &image)?;

    // Scale the image to fit the entire cell.
    image = image.set_scale_to_size(200, 140, false);
    worksheet.write_with_format(2, 0, "Image scaled to fit cell:", &center)?;
    worksheet.insert_image(2, 1, &image)?;

    // Scale the image to fit the defined size region while maintaining the
    // aspect ratio. In this case it is scaled to the smaller of the width or
    // height scales.
    image = image.set_scale_to_size(200, 140, true);
    worksheet.write_with_format(4, 0, "Image scaled with a fixed aspect ratio:", &center)?;
    worksheet.insert_image(4, 1, &image)?;

    // Save the file to disk.
    workbook.save("image.xlsx")?;

    Ok(())
}
