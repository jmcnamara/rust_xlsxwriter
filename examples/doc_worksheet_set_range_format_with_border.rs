// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the format for a range of
//! worksheet cells and also adding a border.

use rust_xlsxwriter::{Format, FormatBorder, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Add some formats.
    let inner_border = Format::new().set_border(FormatBorder::Thin);
    let outer_border = Format::new().set_border(FormatBorder::Double);

    // Write an array of data.
    let data = [
        [10, 11, 12, 13, 14],
        [20, 21, 22, 23, 24],
        [30, 31, 32, 33, 34],
    ];
    worksheet.write_row_matrix(1, 1, data)?;

    // Add formatting to the cells.
    worksheet.set_range_format_with_border(1, 1, 3, 5, &inner_border, &outer_border)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
