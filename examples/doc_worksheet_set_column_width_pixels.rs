// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the width of columns in Excel in
//! pixels.
use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new("worksheet.xlsx");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add some text.
    worksheet.write_string_only(0, 0, "Normal")?;
    worksheet.write_string_only(0, 2, "Wider")?;
    worksheet.write_string_only(0, 4, "Narrower")?;

    // Set the column width in pixels.
    worksheet.set_column_width_pixels(2, 2, 117)?; // Single column.
    worksheet.set_column_width_pixels(4, 5, 33)?; // Column range.

    workbook.close()?;

    Ok(())
}
