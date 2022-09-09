// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting cell diagonal borders.

use rust_xlsxwriter::{Format, Workbook, XlsxBorder, XlsxColor, XlsxDiagonalBorder, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new("formats.xlsx");
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new()
        .set_border_diagonal(XlsxBorder::Thin)
        .set_border_diagonal_type(XlsxDiagonalBorder::BorderUp);

    let format2 = Format::new()
        .set_border_diagonal(XlsxBorder::Thin)
        .set_border_diagonal_type(XlsxDiagonalBorder::BorderDown);

    let format3 = Format::new()
        .set_border_diagonal(XlsxBorder::Thin)
        .set_border_diagonal_type(XlsxDiagonalBorder::BorderUpDown);

    let format4 = Format::new()
        .set_border_diagonal(XlsxBorder::Thin)
        .set_border_diagonal_type(XlsxDiagonalBorder::BorderUpDown)
        .set_border_diagonal_color(XlsxColor::Red);

    worksheet.write_blank(1, 1, &format1)?;
    worksheet.write_blank(3, 1, &format2)?;
    worksheet.write_blank(5, 1, &format3)?;
    worksheet.write_blank(7, 1, &format4)?;

    workbook.close()?;

    Ok(())
}
