// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the cell pattern (with colors).

use rust_xlsxwriter::{Format, FormatPattern, Workbook, XlsxColor, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new()
        .set_background_color(XlsxColor::Green)
        .set_pattern(FormatPattern::Solid);

    let format2 = Format::new()
        .set_background_color(XlsxColor::Yellow)
        .set_foreground_color(XlsxColor::Red)
        .set_pattern(FormatPattern::DarkVertical);

    worksheet.write_string(0, 0, "Rust", &format1)?;
    worksheet.write_blank(1, 0, &format2)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
