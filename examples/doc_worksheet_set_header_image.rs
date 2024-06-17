// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates adding a header image to a worksheet.

use rust_xlsxwriter::{HeaderImagePosition, Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Scale the image so it fits in the header.
    let image = Image::new("examples/rust_logo.png")?
        .set_scale_height(0.5)
        .set_scale_width(0.5);

    // Insert the watermark image in the header.
    worksheet.set_header("&C&[Picture]");
    worksheet.set_header_image(&image, HeaderImagePosition::Center)?;

    // Increase the top margin to 1.2 for clarity. The -1.0 values are ignored.
    worksheet.set_margins(-1.0, -1.0, 1.2, -1.0, -1.0, -1.0);

    // Set Page View mode so the watermark is visible.
    worksheet.set_view_page_layout();

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
