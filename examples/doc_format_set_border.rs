// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting a cell border.

use rust_xlsxwriter::{Format, Workbook, XlsxBorder, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new().set_border(XlsxBorder::Thin);
    let format2 = Format::new().set_border(XlsxBorder::Dotted);
    let format3 = Format::new().set_border(XlsxBorder::Double);

    worksheet.write_blank(1, 1, &format1)?;
    worksheet.write_blank(3, 1, &format2)?;
    worksheet.write_blank(5, 1, &format3)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
