// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates writing an array of data as a row to a
//! worksheet.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Some array data to write.
    let data = [1, 2, 3, 4, 5];

    // Write the array data as a row.
    worksheet.write_row(0, 0, data)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
