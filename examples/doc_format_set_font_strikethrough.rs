// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the text strikethrough property
//! for a format.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let format = Format::new().set_font_strikethrough();

    worksheet.write_string(0, 0, "Strikethrough Text", &format)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
