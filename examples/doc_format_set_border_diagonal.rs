// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting cell diagonal borders.

use rust_xlsxwriter::{Format, FormatBorder, FormatDiagonalBorder, Workbook, XlsxColor, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new()
        .set_border_diagonal(FormatBorder::Thin)
        .set_border_diagonal_type(FormatDiagonalBorder::BorderUp);

    let format2 = Format::new()
        .set_border_diagonal(FormatBorder::Thin)
        .set_border_diagonal_type(FormatDiagonalBorder::BorderDown);

    let format3 = Format::new()
        .set_border_diagonal(FormatBorder::Thin)
        .set_border_diagonal_type(FormatDiagonalBorder::BorderUpDown);

    let format4 = Format::new()
        .set_border_diagonal(FormatBorder::Thin)
        .set_border_diagonal_type(FormatDiagonalBorder::BorderUpDown)
        .set_border_diagonal_color(XlsxColor::Red);

    worksheet.write_blank(1, 1, &format1)?;
    worksheet.write_blank(3, 1, &format2)?;
    worksheet.write_blank(5, 1, &format3)?;
    worksheet.write_blank(7, 1, &format4)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
