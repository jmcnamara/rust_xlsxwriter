// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the text reading direction. This
//! is useful when creating Arabic, Hebrew or other near or far eastern
//! worksheets.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_column_width(0, 25)?;

    let format1 = Format::new().set_reading_direction(1);
    let format2 = Format::new().set_reading_direction(2);

    worksheet.write_string(0, 0, "نص عربي / English text")?;
    worksheet.write_string_with_format(1, 0, "نص عربي / English text", &format1)?;
    worksheet.write_string_with_format(2, 0, "نص عربي / English text", &format2)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
