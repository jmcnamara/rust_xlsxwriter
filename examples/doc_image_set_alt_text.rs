// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! This example shows how to create an image object and set the alternative
//! text to help accessibility.

use rust_xlsxwriter::{Image, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Create a new image object and set the alternative text.
    let image = Image::new("examples/rust_logo.png")?.set_alt_text(
        "A circular logo with gear teeth on the outside \
        and a large letter R on the inside.\n\n\
        The logo of the Rust programming language.",
    );

    // Insert the image.
    worksheet.insert_image(1, 2, &image)?;

    // Save the file to disk.
    workbook.save("image.xlsx")?;

    Ok(())
}
