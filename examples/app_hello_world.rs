// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

//! Create a simple Hello World style Excel spreadsheet using the
//! rust_xlsxwriter library.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file.
    let mut workbook = Workbook::new("hello.xlsx");

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write a string to cell (0, 0) = A1.
    worksheet.write_string_only(0, 0, "Hello")?;

    // Write a number to cell (1, 0) = A2.
    worksheet.write_number_only(1, 0, 12345)?;

    // Close the file.
    workbook.close()?;

    Ok(())
}
