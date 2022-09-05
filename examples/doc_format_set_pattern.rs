// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the cell pattern (with colors).

use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError, XlsxPattern};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new("formats.xlsx");

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new()
        .set_background_color(XlsxColor::Green)
        .set_pattern(XlsxPattern::Solid);

    let format2 = Format::new()
        .set_background_color(XlsxColor::Yellow)
        .set_foreground_color(XlsxColor::Red)
        .set_pattern(XlsxPattern::DarkVertical);

    worksheet.write_string(0, 0, "Rust", &format1)?;
    worksheet.write_blank(1, 0, &format2)?;

    workbook.close()?;

    Ok(())
}
