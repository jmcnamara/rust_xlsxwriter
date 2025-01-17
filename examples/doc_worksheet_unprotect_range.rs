// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates unprotecting ranges in a protected
//! worksheet.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Protect the worksheet from modification.
    worksheet.protect();

    // Unprotect range D4:F10.
    worksheet.unprotect_range(4, 3, 9, 5)?;

    // Unprotect single cell B3 by repeating (row, col).
    worksheet.unprotect_range(2, 1, 2, 1)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
