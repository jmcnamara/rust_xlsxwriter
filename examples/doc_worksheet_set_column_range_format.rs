// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the format for all the columns in
//! an Excel worksheet. This effectively, and efficiently, sets the format for
//! the entire worksheet.
use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add a format.
    let cell_format = Format::new()
        .set_background_color("#F9F2EC")
        .set_border(FormatBorder::Thin);

    // Set the column format for the entire worksheet.
    worksheet.set_column_range_format(0, 16_383, &cell_format)?;

    // Add some unformatted text that adopts the column format.
    worksheet.write_string(1, 1, "Hello")?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
