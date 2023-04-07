// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates using different `XlsxColor` enum values to
//! set the color of some text in a worksheet.

use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_column_width(0, 14)?;

    let format1 = Format::new().set_font_color(XlsxColor::Red);
    let format2 = Format::new().set_font_color(XlsxColor::Green);
    let format3 = Format::new().set_font_color(XlsxColor::RGB(0x4F_02_6A));
    let format4 = Format::new().set_font_color(XlsxColor::RGB(0x73_CC_5F));
    let format5 = Format::new().set_font_color(XlsxColor::Theme(4, 0));
    let format6 = Format::new().set_font_color(XlsxColor::Theme(9, 4));

    worksheet.write_string_with_format(0, 0, "Red", &format1)?;
    worksheet.write_string_with_format(1, 0, "Green", &format2)?;
    worksheet.write_string_with_format(2, 0, "#4F026A", &format3)?;
    worksheet.write_string_with_format(3, 0, "#73CC5F", &format4)?;
    worksheet.write_string_with_format(4, 0, "Theme (4, 0)", &format5)?;
    worksheet.write_string_with_format(5, 0, "Theme (9, 4)", &format6)?;

    workbook.save("colors.xlsx")?;

    Ok(())
}
