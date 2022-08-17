// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates using different XlsxColor enum values to
//! set the color of some text in a worksheet.

use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new("colors.xlsx");

    let format1 = Format::new().set_font_color(XlsxColor::Red);
    let format2 = Format::new().set_font_color(XlsxColor::Green);
    let format3 = Format::new().set_font_color(XlsxColor::RGB(0x4F026A));
    let format4 = Format::new().set_font_color(XlsxColor::RGB(0x73CC5F));
    let format5 = Format::new().set_font_color(XlsxColor::RGB(0xFFACFF));
    let format6 = Format::new().set_font_color(XlsxColor::RGB(0xCC7E16));

    let worksheet = workbook.add_worksheet();
    worksheet.write_string(0, 0, "Red", &format1)?;
    worksheet.write_string(1, 0, "Green", &format2)?;
    worksheet.write_string(2, 0, "#4F026A", &format3)?;
    worksheet.write_string(3, 0, "#73CC5F", &format4)?;
    worksheet.write_string(4, 0, "#FFACFF", &format5)?;
    worksheet.write_string(5, 0, "#CC7E16", &format6)?;

    workbook.close()?;

    Ok(())
}
