// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates protecting a worksheet from editing with
//! a password.

use rust_xlsxwriter::{Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();

    // Protect the worksheet from modification.
    worksheet.protect_with_password("abc123");

    worksheet.write_string(0, 0, "Unlock the worksheet to edit the cell")?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
