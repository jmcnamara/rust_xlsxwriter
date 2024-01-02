// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing a blank cell with formatting,
//! i.e., a cell that has no data but does have formatting.

use rust_xlsxwriter::{Color, Format, FormatBorder, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new().set_background_color(Color::Yellow);

    let format2 = Format::new()
        .set_background_color(Color::Yellow)
        .set_border(FormatBorder::Thin);

    worksheet.write_blank(1, 1, &format1)?;
    worksheet.write_blank(3, 1, &format2)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
