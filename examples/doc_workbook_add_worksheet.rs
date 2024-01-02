// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates creating adding worksheets to a workbook.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet(); // Sheet1
    worksheet.write_string(0, 0, "Hello")?;

    let worksheet = workbook.add_worksheet().set_name("Foglio2")?;
    worksheet.write_string(0, 0, "Hello")?;

    let worksheet = workbook.add_worksheet(); // Sheet3
    worksheet.write_string(0, 0, "Hello")?;

    workbook.save("workbook.xlsx")?;

    Ok(())
}
