// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the cell background color, with a
//! default solid pattern.

use rust_xlsxwriter::{Format, Workbook, XlsxColor, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new("formats.xlsx");

    // Add a worksheet.
    let worksheet = workbook.add_worksheet();

    let format1 = Format::new().set_background_color(XlsxColor::Green);

    worksheet.write_string(0, 0, "Rust", &format1)?;

    workbook.close()?;

    Ok(())
}
