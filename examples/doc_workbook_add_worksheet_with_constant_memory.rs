// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates adding worksheets in "standard" and
//! "constant memory" modes.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Add a worksheet in standard mode.
    let worksheet = workbook.add_worksheet();
    worksheet.write(0, 0, "Standard")?;

    // Add a worksheet in "constant memory" mode.
    let worksheet = workbook.add_worksheet_with_constant_memory();
    worksheet.write(0, 0, "Constant memory")?;

    workbook.save("workbook.xlsx")?;

    Ok(())
}
