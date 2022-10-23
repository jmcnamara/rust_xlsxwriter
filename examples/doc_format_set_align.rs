// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting various cell alignment
//! properties.

use rust_xlsxwriter::{Format, Workbook, XlsxAlign, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new("formats.xlsx");
    let worksheet = workbook.add_worksheet();

    // Widen the rows/column for clarity.
    worksheet.set_row_height(1, 30)?;
    worksheet.set_row_height(2, 30)?;
    worksheet.set_row_height(3, 30)?;
    worksheet.set_column_width(0, 18)?;

    // Create some alignment formats.
    let format1 = Format::new().set_align(XlsxAlign::Center);

    let format2 = Format::new()
        .set_align(XlsxAlign::Top)
        .set_align(XlsxAlign::Left);

    let format3 = Format::new()
        .set_align(XlsxAlign::VerticalCenter)
        .set_align(XlsxAlign::Center);

    let format4 = Format::new()
        .set_align(XlsxAlign::Bottom)
        .set_align(XlsxAlign::Right);

    worksheet.write_string(0, 0, "Center", &format1)?;
    worksheet.write_string(1, 0, "Top - Left", &format2)?;
    worksheet.write_string(2, 0, "Center - Center", &format3)?;
    worksheet.write_string(3, 0, "Bottom - Right", &format4)?;

    workbook.close()?;

    Ok(())
}
