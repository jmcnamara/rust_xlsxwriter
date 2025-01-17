// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the font name/type for a format.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let format = Format::new().set_font_name("Avenir Black Oblique");

    worksheet.write_string_with_format(0, 0, "Avenir Black Oblique", &format)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
