// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates clearing the formatting from some
//! previously written cells in a worksheet.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add a format.
    let format = Format::new().set_font_color("#FF0000");

    // Some array data to write.
    let data = [
        [10, 11, 12, 13, 14],
        [20, 21, 22, 23, 24],
        [30, 31, 32, 33, 34],
    ];

    // Write the array data as a series of rows.
    worksheet.write_row_with_format(0, 0, data[0], &format)?;
    worksheet.write_row_with_format(1, 0, data[1], &format)?;
    worksheet.write_row_with_format(2, 0, data[2], &format)?;

    // Clear the format from the first and last cells in the data.
    worksheet.clear_cell_format(0, 0);
    worksheet.clear_cell_format(2, 4);

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
