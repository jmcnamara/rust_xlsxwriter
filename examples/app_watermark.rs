// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of adding a worksheet watermark image using the `rust_xlsxwriter`
//! library. This is based on the method of putting an image in the worksheet
//! header as suggested in the Microsoft documentation.

use rust_xlsxwriter::{HeaderImagePosition, Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    let image = Image::new("examples/watermark.png")?;

    // Insert the watermark image in the header.
    worksheet.set_header("&C&[Picture]");
    worksheet.set_header_image(&image, HeaderImagePosition::Center)?;

    // Set Page View mode so the watermark is visible.
    worksheet.set_view_page_layout();

    // Save the file to disk.
    workbook.save("watermark.xlsx")?;

    Ok(())
}
