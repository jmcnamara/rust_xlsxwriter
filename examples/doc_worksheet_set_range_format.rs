// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the format of worksheet cells
//! separately from writing the cell data.

use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add a format.
    let border = Format::new().set_border(FormatBorder::Thin);

    // Write an array of data.
    let data = [
        [10, 11, 12, 13, 14],
        [20, 21, 22, 23, 24],
        [30, 31, 32, 33, 34],
    ];
    worksheet.write_row_matrix(1, 1, data)?;

    // Add formatting to the cells.
    worksheet.set_range_format(1, 1, 3, 5, &border)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
