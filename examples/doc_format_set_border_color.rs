// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting a cell border and color.

use rust_xlsxwriter::{Color, Format, FormatBorder, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new()
        .set_border(FormatBorder::Thin)
        .set_border_color(Color::Blue);

    let format2 = Format::new()
        .set_border(FormatBorder::Dotted)
        .set_border_color(Color::Red);

    let format3 = Format::new()
        .set_border(FormatBorder::Double)
        .set_border_color(Color::Green);

    worksheet.write_blank(1, 1, &format1)?;
    worksheet.write_blank(3, 1, &format2)?;
    worksheet.write_blank(5, 1, &format3)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
