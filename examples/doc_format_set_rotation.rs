// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting text rotation for a cell.

use rust_xlsxwriter::{Format, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Widen the rows/column for clarity.
    worksheet.set_row_height(0, 30)?;
    worksheet.set_row_height(1, 30)?;
    worksheet.set_row_height(2, 60)?;

    // Create some alignment formats.
    let format1 = Format::new().set_rotation(30);
    let format2 = Format::new().set_rotation(-30);
    let format3 = Format::new().set_rotation(270);

    worksheet.write_string_with_format(0, 0, "Rust", &format1)?;
    worksheet.write_string_with_format(1, 0, "Rust", &format2)?;
    worksheet.write_string_with_format(2, 0, "Rust", &format3)?;

    workbook.save("formats.xlsx")?;

    Ok(())
}
