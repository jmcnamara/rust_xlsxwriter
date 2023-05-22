// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the foreground/pattern color.

use rust_xlsxwriter::{Color, Format, FormatPattern, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new()
        .set_background_color(Color::Yellow)
        .set_foreground_color(Color::Red)
        .set_pattern(FormatPattern::DarkVertical);

    worksheet.write_blank(0, 0, &format1)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
