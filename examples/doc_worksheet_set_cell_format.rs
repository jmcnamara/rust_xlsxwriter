// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the format of a worksheet cell
//! separately from writing the cell data.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add some formats.
    let red = Format::new().set_font_color("#FF0000");
    let green = Format::new().set_font_color("#00FF00");

    // Write an array of data.
    let data = [
        [10, 11, 12, 13, 14],
        [20, 21, 22, 23, 24],
        [30, 31, 32, 33, 34],
    ];
    worksheet.write_row_matrix(1, 1, data)?;

    // Add formatting to some of the cells.
    worksheet.set_cell_format(1, 1, &red)?;
    worksheet.set_cell_format(2, 3, &green)?;
    worksheet.set_cell_format(3, 5, &red)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
