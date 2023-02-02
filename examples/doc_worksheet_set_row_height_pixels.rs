// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the height for a row in Excel.
use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add some text.
    worksheet.write_string(0, 0, "Normal")?;
    worksheet.write_string(2, 0, "Taller")?;

    // Set the row height in pixels.
    worksheet.set_row_height_pixels(2, 40)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
