// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the format of worksheet cells
//! when writing the cell data.

use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add a format.
    let border = Format::new().set_border(FormatBorder::Thin);

    // Some data to write.
    let data = [
        [10, 11, 12, 13, 14],
        [20, 21, 22, 23, 24],
        [30, 31, 32, 33, 34],
    ];

    // Write the data with formatting.
    for (row_num, col) in data.iter().enumerate() {
        for (col_num, cell) in col.iter().enumerate() {
            let row_num = row_num as u32 + 1;
            let col_num = col_num as u16 + 1;
            worksheet.write_with_format(row_num, col_num, *cell, &border)?;
        }
    }

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
