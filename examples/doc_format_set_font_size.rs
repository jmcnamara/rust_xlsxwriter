// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the font size for a format.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new("formats.xlsx");
    let worksheet = workbook.add_worksheet();

    let format = Format::new().set_font_size(30);

    worksheet.write_string(0, 0, "Font Size 30", &format)?;

    workbook.close()?;

    Ok(())
}
