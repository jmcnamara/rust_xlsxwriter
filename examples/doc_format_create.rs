// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates create a new format and setting the
//! properties.

use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new("formats.xlsx");

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    // Create a new format and set some properties.
    let format = Format::new()
        .set_bold()
        .set_italic()
        .set_font_color(XlsxColor::Red);

    worksheet.write_string(0, 0, "Hello", &format)?;

    workbook.close()?;

    Ok(())
}
