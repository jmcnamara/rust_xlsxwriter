// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates getting worksheet reference by index.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Start with a reference to worksheet1.
    let mut worksheet1 = workbook.add_worksheet();
    worksheet1.write_string(0, 0, "Hello")?;

    // If we don't try to use the workbook1 reference again we can switch to
    // using a reference to worksheet2.
    let mut worksheet2 = workbook.add_worksheet();
    worksheet2.write_string(0, 0, "Hello")?;

    // Stop using worksheet2 and move back to worksheet1.
    worksheet1 = workbook.worksheet_from_index(0)?;
    worksheet1.write_string(1, 0, "Sheet1")?;

    // Stop using worksheet1 and move back to worksheet2.
    worksheet2 = workbook.worksheet_from_index(1)?;
    worksheet2.write_string(1, 0, "Sheet2")?;

    workbook.save("workbook.xlsx")?;

    Ok(())
}
