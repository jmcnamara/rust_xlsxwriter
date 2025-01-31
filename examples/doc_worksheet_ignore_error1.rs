// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! This example demonstrates an Excel cell warning.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write a string that looks like a number.
    worksheet.write_string(1, 2, "123")?;

    // Save the file to disk.
    workbook.save("worksheet.xlsx")?;

    Ok(())
}
