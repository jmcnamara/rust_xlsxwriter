// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating new worksheet objects and then
//! adding them to a workbook.

use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new workbook.
    let mut workbook = Workbook::new();

    // Create new worksheets.
    let mut worksheet1 = Worksheet::new();
    let mut worksheet2 = Worksheet::new();

    // Use the first workbook.
    worksheet1.write_string(0, 0, "Hello")?;
    worksheet1.write_string(1, 0, "Sheet1")?;

    // Use the second workbook.
    worksheet2.write_string(0, 0, "Hello")?;
    worksheet2.write_string(1, 0, "Sheet2")?;

    // Add the worksheets to the workbook.
    workbook.push_worksheet(worksheet1);
    workbook.push_worksheet(worksheet2);

    // Save the workbook.
    workbook.save("worksheets.xlsx")?;

    Ok(())
}
