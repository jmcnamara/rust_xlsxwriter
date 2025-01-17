// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the columns to repeat on each
//! printed page.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet1 = workbook.add_worksheet();
    // Repeat the first column in the printed output.
    worksheet1.set_repeat_columns(0, 0)?;

    let worksheet2 = workbook.add_worksheet();
    // Repeat the first 2 columns in the printed output.
    worksheet2.set_repeat_columns(0, 1)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
