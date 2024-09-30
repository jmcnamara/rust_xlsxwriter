// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates adding worksheets in "standard" and
//! "low memory" modes.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet in standard mode.
    let worksheet = workbook.add_worksheet();
    worksheet.write(0, 0, "Standard")?;

    // Add a worksheet in "low memory" mode.
    let worksheet = workbook.add_worksheet_with_low_memory();
    worksheet.write(0, 0, "Low memory")?;

    workbook.save("workbook.xlsx")?;

    Ok(())
}
