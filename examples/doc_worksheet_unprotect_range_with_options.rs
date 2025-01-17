// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates unprotecting ranges in a protected
//! worksheet, with additional options.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Protect the worksheet from modification.
    worksheet.protect();

    // Unprotect range D4:F10 and give it a user defined name.
    worksheet.unprotect_range_with_options(4, 3, 9, 5, "MyRange", "")?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
