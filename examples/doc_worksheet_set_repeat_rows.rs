// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting the rows to repeat on each
//! printed page.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    let worksheet1 = workbook.add_worksheet();
    // Repeat the first row in the printed output.
    worksheet1.set_repeat_rows(0, 0)?;

    let worksheet2 = workbook.add_worksheet();
    // Repeat the first 2 rows in the printed output.
    worksheet2.set_repeat_rows(0, 1)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
