// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the cell pattern (with colors).

use rust_xlsxwriter::{Color, Format, FormatPattern, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new()
        .set_background_color(Color::Green)
        .set_pattern(FormatPattern::Solid);

    let format2 = Format::new()
        .set_background_color(Color::Yellow)
        .set_foreground_color(Color::Red)
        .set_pattern(FormatPattern::DarkVertical);

    worksheet.write_with_format(0, 0, "Rust", &format1)?;
    worksheet.write_blank(1, 0, &format2)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
