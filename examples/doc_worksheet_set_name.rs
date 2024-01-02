// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting user defined worksheet names
//! and the default values when a name isn't set.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let _worksheet1 = workbook.add_worksheet(); // Defaults to Sheet1
    let _worksheet2 = workbook.add_worksheet().set_name("Foglio2");
    let _worksheet3 = workbook.add_worksheet().set_name("Data");
    let _worksheet4 = workbook.add_worksheet(); // Defaults to Sheet4

    workbook.save("worksheets.xlsx")?;

    Ok(())
}
