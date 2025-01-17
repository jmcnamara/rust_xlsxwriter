// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates getting a worksheet name.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Try name() using a default sheet name.
    let worksheet = workbook.add_worksheet();
    assert_eq!("Sheet1", worksheet.name());

    // Try name() using a user defined sheet name.
    let worksheet = workbook.add_worksheet().set_name("Data")?;
    assert_eq!("Data", worksheet.name());

    workbook.save("workbook.xlsx")?;

    Ok(())
}
