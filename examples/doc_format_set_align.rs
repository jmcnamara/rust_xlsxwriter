// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting various cell alignment
//! properties.

use rust_xlsxwriter::{Format, FormatAlign, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Widen the rows/column for clarity.
    worksheet.set_row_height(1, 30)?;
    worksheet.set_row_height(2, 30)?;
    worksheet.set_row_height(3, 30)?;
    worksheet.set_column_width(0, 18)?;

    // Create some alignment formats.
    let format1 = Format::new().set_align(FormatAlign::Center);

    let format2 = Format::new()
        .set_align(FormatAlign::Top)
        .set_align(FormatAlign::Left);

    let format3 = Format::new()
        .set_align(FormatAlign::VerticalCenter)
        .set_align(FormatAlign::Center);

    let format4 = Format::new()
        .set_align(FormatAlign::Bottom)
        .set_align(FormatAlign::Right);

    worksheet.write_string_with_format(0, 0, "Center", &format1)?;
    worksheet.write_string_with_format(1, 0, "Top - Left", &format2)?;
    worksheet.write_string_with_format(2, 0, "Center - Center", &format3)?;
    worksheet.write_string_with_format(3, 0, "Bottom - Right", &format4)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
