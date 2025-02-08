// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a format that is a
//! combination of two formats.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    // Add a format. Green fill with dark green text.
    let format1 = Format::new()
        .set_font_color("006100")
        .set_background_color("C6EFCE");

    // Add a format. Bold and italic.
    let format2 = Format::new().set_bold().set_italic();

    // Create a new format based on a merge of two formats.
    let merged = format1.merge(&format2);

    // Write some strings with the formats.
    worksheet.write_string_with_format(0, 0, "Format 1", &format1)?;
    worksheet.write_string_with_format(2, 0, "Format 2", &format2)?;
    worksheet.write_string_with_format(4, 0, "Merged", &merged)?;

    // Save the file.
    workbook.save("formats.xlsx")?;

    Ok(())
}
