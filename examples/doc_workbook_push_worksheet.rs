// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating a standalone worksheet object
//! and then adding it to a workbook.

use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let mut worksheet = Worksheet::new();

    // Use the worksheet object.
    worksheet.write_string(0, 0, "Hello")?;

    // Add it to the workbook.
    workbook.push_worksheet(worksheet);

    // Save the workbook.
    workbook.save("workbook.xlsx")?;

    Ok(())
}
