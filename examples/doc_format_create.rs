// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates create a new format and setting the
//! properties.

use rust_xlsxwriter::{Color, Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    // Create a new format and set some properties.
    let format = Format::new()
        .set_bold()
        .set_italic()
        .set_font_color(Color::Red);

    worksheet.write_with_format(0, 0, "Hello", &format)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
