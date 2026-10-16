// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates hiding zero values in worksheet cells.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Write some numbers, including a zero.
    worksheet.write(0, 0, 0)?;
    worksheet.write(1, 0, 1)?;
    worksheet.write(2, 0, 2)?;
    worksheet.write(3, 0, 3)?;

    // Hide the zero values.
    worksheet.set_hidden_zeroes(true);

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
