// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the text strikethrough property
//! for a format.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new("formats.xlsx");
    let worksheet = workbook.add_worksheet();

    let format = Format::new().set_font_strikethrough();

    worksheet.write_string(0, 0, "Strikethrough Text", &format)?;

    workbook.close()?;

    Ok(())
}
