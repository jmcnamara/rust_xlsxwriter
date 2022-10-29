// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the format for a row in Excel.
use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add for formats.
    let bold_format = Format::new().set_bold();
    let red_format = Format::new().set_font_color(XlsxColor::Red);

    // Set the row format.
    worksheet.set_row_format(1, &red_format)?;

    // Add some unformatted text that adopts the row format.
    worksheet.write_string_only(1, 0, "Hello")?;

    // Add some formatted text that overrides the row format.
    worksheet.write_string(1, 2, "Hello", &bold_format)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
